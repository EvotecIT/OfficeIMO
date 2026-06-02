using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using SdtContentPicture = DocumentFormat.OpenXml.Wordprocessing.SdtContentPicture;
using TabStop = DocumentFormat.OpenXml.Wordprocessing.TabStop;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using V = DocumentFormat.OpenXml.Vml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Gets the first image associated with this run, if any.
        /// </summary>
        public WordImage? Image {
            get {
                if (_run != null) {
                    // DrawingML pictures
                    var drawing = _run.ChildElements.OfType<WordDrawing>().FirstOrDefault();
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
        /// Gets the chart contained in this paragraph, if present.
        /// </summary>
        public WordChart? Chart {
            get {
                if (_run is not null) {
                    var drawing = _run.ChildElements.OfType<WordDrawing>().FirstOrDefault();
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
                    var drawing = _run.ChildElements.OfType<WordDrawing>().FirstOrDefault();
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
        /// Gets a value indicating whether an image is found in the paragraph.
        /// </summary>
        public bool IsImage => Image is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph has an embedded object.
        /// </summary>
        public bool IsEmbeddedObject => EmbeddedObject is not null;
        /// <summary>
        /// Gets a value indicating whether a chart is associated with the paragraph.
        /// </summary>
        public bool IsChart => Chart is not null;

        /// <summary>
        /// Gets a value indicating whether SmartArt is present in the paragraph.
        /// </summary>
        public bool IsSmartArt => SmartArt is not null;
        /// <summary>
        /// Gets the <see cref="WordTextBox"/> contained within the paragraph, if any.
        /// </summary>
        public WordTextBox? TextBox {
            get {
                if (_run is not null) {
                    // DrawingML text boxes
                    var drawing = _run.ChildElements.OfType<WordDrawing>().FirstOrDefault();
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
                    var drawing = _run.ChildElements.OfType<WordDrawing>().FirstOrDefault();
                    if (drawing is null) {
                        foreach (var ac in _run.ChildElements.OfType<AlternateContent>()) {
                            var choice = ac.ChildElements.OfType<AlternateContentChoice>().FirstOrDefault();
                            if (choice is not null) {
                                drawing = choice.Descendants<WordDrawing>().FirstOrDefault();
                                if (drawing is not null) {
                                    break;
                                }
                            }
                            var fallback = ac.ChildElements.OfType<AlternateContentFallback>().FirstOrDefault();
                            if (fallback is not null) {
                                drawing = fallback.Descendants<WordDrawing>().FirstOrDefault();
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
