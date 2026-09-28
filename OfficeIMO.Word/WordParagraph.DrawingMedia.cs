using DocumentFormat.OpenXml;
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
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Gets the first image associated with this run, if any.
        /// </summary>
        public WordImage? Image {
            get => EnumerateImages().FirstOrDefault();
        }

        /// <summary>Enumerates every DrawingML or VML image represented by this run.</summary>
        internal IEnumerable<WordImage> EnumerateImages() {
            if (_run == null) yield break;

            foreach (WordDrawing drawing in EnumerateEffectiveRunContent().SelectMany(
                         element => element is WordDrawing direct
                             ? new[] { direct }
                             : element.Descendants<WordDrawing>())) {
                bool inlinePicture = drawing.Inline?.Graphic?.GraphicData?.ChildElements
                    .OfType<DocumentFormat.OpenXml.Drawing.Pictures.Picture>()
                    .Any() == true;
                bool anchoredPicture = drawing.Anchor?.Elements<Graphic>()
                    .Any(graphic => graphic.GraphicData?.ChildElements
                        .OfType<DocumentFormat.OpenXml.Drawing.Pictures.Picture>()
                        .Any() == true) == true;
                if (inlinePicture || anchoredPicture) {
                    yield return new WordImage(_document, drawing);
                }
            }

            foreach (V.Shape shape in EnumerateEffectiveRunContent().SelectMany(
                         element => element is V.Shape direct
                             ? new[] { direct }
                             : element.Descendants<V.Shape>())) {
                if (shape.GetFirstChild<V.ImageData>() != null) {
                    yield return new WordImage(_document, _paragraph, _run, shape);
                }
            }
        }

        /// <summary>
        /// Enumerates direct run content while selecting the active markup-compatibility branch.
        /// </summary>
        private IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> EnumerateEffectiveRunContent() {
            foreach (OpenXmlElement child in VisibleSourceRunChildren()) {
                if (child is not AlternateContent alternateContent) {
                    yield return child;
                    continue;
                }

                DocumentFormat.OpenXml.OpenXmlCompositeElement? branch =
                    WordAlternateContentResolver.SelectBranch(alternateContent);
                if (branch == null) continue;
                foreach (DocumentFormat.OpenXml.OpenXmlElement branchChild in branch.ChildElements) {
                    yield return branchChild;
                }
            }
        }

        private IEnumerable<OpenXmlElement> VisibleSourceRunChildren() =>
            _visibleRunSourceChildren ?? (IEnumerable<OpenXmlElement>?)_run?.ChildElements ?? Array.Empty<OpenXmlElement>();

        private IEnumerable<T> VisibleSourceRunDescendants<T>() where T : OpenXmlElement =>
            VisibleSourceRunChildren().SelectMany(child =>
                child is T direct ? new[] { direct } : child.Descendants<T>());

        /// <summary>
        /// Gets the embedded object associated with this run, if any.
        /// </summary>
        public WordEmbeddedObject? EmbeddedObject {
            get {
                if (_run != null) {
                    var ole = VisibleSourceRunDescendants<Ovml.OleObject>().FirstOrDefault();
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
                    foreach (WordDrawing drawing in VisibleSourceRunChildren().OfType<WordDrawing>()) {
                        if (drawing.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any()) {
                            return new WordChart(_document, this, drawing);
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
                    var drawing = VisibleSourceRunChildren().OfType<WordDrawing>().FirstOrDefault();
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
                    var drawing = VisibleSourceRunChildren().OfType<WordDrawing>().FirstOrDefault();
                    if (drawing is not null) {
                        if (drawing.Descendants<Wps.TextBoxInfo2>().Any()) {
                            return new WordTextBox(_document, _paragraph, _run, selectedDrawing: drawing);
                        }
                    }

                    // Legacy text boxes wrapped in AlternateContent (Word 2007)
                    bool choiceHasOnlyShape = false;
                    foreach (var ac in VisibleSourceRunChildren().OfType<AlternateContent>()) {
                        DocumentFormat.OpenXml.OpenXmlCompositeElement? branch =
                            WordAlternateContentResolver.SelectBranch(ac);
                        if (branch is not null) {
                            bool branchHasTextBox = branch.Descendants<Wps.TextBoxInfo2>().Any() || branch.Descendants<V.TextBox>().Any();
                            if (branchHasTextBox) {
                                return new WordTextBox(_document, _paragraph, _run,
                                    selectedAlternateContent: ac,
                                    selectedVmlTextBox: branch.Descendants<V.TextBox>().FirstOrDefault());
                            }
                            bool hasShape = branch.Descendants<Wps.WordprocessingShape>().Any() ||
                                branch.Descendants<V.Shape>().Any(s => !s.Descendants<V.ImageData>().Any() && !s.Descendants<V.TextBox>().Any());
                            if (hasShape) {
                                choiceHasOnlyShape = true;
                                continue;
                            }
                        }
                    }
                    if (choiceHasOnlyShape) {
                        return null;
                    }

                    // VML text boxes
                    if (VisibleSourceRunDescendants<V.TextBox>().FirstOrDefault() is { } vmlTextBox) {
                        return new WordTextBox(_document, _paragraph, _run, selectedVmlTextBox: vmlTextBox);
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
                    OpenXmlElement? vmlShape = VisibleSourceRunChildren()
                        .SelectMany(child => child.Descendants().Prepend(child))
                        .FirstOrDefault(element => element is V.Rectangle or V.RoundRectangle or V.Oval or V.Line or V.PolyLine ||
                            element is V.Shape shape && !shape.Descendants<V.ImageData>().Any() && !shape.Descendants<V.TextBox>().Any());
                    if (vmlShape != null)
                        return new WordShape(_document, _paragraph, _run, selectedVmlShape: vmlShape);

                    // DrawingML shapes (non-pictures and not text boxes)
                    var drawing = VisibleSourceRunChildren().OfType<WordDrawing>().FirstOrDefault();
                    if (drawing is null) {
                        foreach (var ac in VisibleSourceRunChildren().OfType<AlternateContent>()) {
                            DocumentFormat.OpenXml.OpenXmlCompositeElement? branch =
                                WordAlternateContentResolver.SelectBranch(ac);
                            drawing = branch?.Descendants<WordDrawing>().FirstOrDefault();
                            if (drawing is not null) break;
                        }
                    }
                    if (drawing is not null) {
                        bool hasPicture = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().Any();
                        bool hasTextBox = drawing.Descendants<Wps.TextBoxInfo2>().Any();
                        bool hasShapeGroup = drawing.Descendants<Wpg.WordprocessingGroup>().Any();
                        bool hasShape = drawing.Descendants<Wps.WordprocessingShape>().Any();
                        if (!hasPicture && !hasTextBox && !hasShapeGroup && hasShape) {
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
                    var line = VisibleSourceRunDescendants<V.Line>().FirstOrDefault();
                    if (line is not null) {
                        return new WordLine(_document, _paragraph, _run, line);
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
