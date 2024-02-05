using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Graphic = DocumentFormat.OpenXml.Drawing.Graphic;

namespace OfficeIMO.Word {
    public class WordTextBox {
        private WordDocument _document;
        private WordParagraph _wordParagraph;
        private readonly WordHeaderFooter _headerFooter;
        private Run _run => _wordParagraph._run;

        /// <summary>
        /// Add a new text box to the document
        /// </summary>
        /// <param name="wordDocument"></param>
        /// <param name="text"></param>
        /// <param name="wrapTextImage"></param>
        public WordTextBox(WordDocument wordDocument, string text, WrapTextImage wrapTextImage) {
            var paragraph = new WordParagraph(wordDocument, true, true);
            wordDocument.AddParagraph(paragraph);
            paragraph._run.Append(new RunProperties());
            AddAlternateContent(wordDocument, paragraph, text, wrapTextImage);

            _document = wordDocument;
            _wordParagraph = paragraph;
        }

        /// <summary>
        /// Initialize a text box from an existing paragraph
        /// </summary>
        /// <param name="wordDocument"></param>
        /// <param name="paragraph"></param>
        /// <param name="run"></param>
        public WordTextBox(WordDocument wordDocument, Paragraph paragraph, Run run) {
            _document = wordDocument;
            _wordParagraph = new WordParagraph(wordDocument, paragraph, run);
        }

        public WordTextBox(WordDocument wordDocument, WordHeaderFooter wordHeaderFooter, string text, WrapTextImage wrapTextImage) {
            _document = wordDocument;
            _headerFooter = wordHeaderFooter;

            var paragraph = wordHeaderFooter.AddParagraph(newRun: true);
            paragraph._run.Append(new RunProperties());
            AddAlternateContent(wordDocument, paragraph, text, wrapTextImage);

            _wordParagraph = paragraph;
        }

        public WordTextBox(WordDocument wordDocument, WordParagraph paragraph, string text, WrapTextImage wrapTextImage) {
            _document = wordDocument;

            paragraph._run.Append(new RunProperties());
            AddAlternateContent(wordDocument, paragraph, text, wrapTextImage);

            _wordParagraph = paragraph;
        }

        /// <summary>
        /// Allows to set the text of the text box
        /// For more advanced text formatting use WordParagraph property
        /// </summary>
        public string Text {
            get {
                if (_sdtBlock != null) {

                    var run = _sdtBlock.GetFirstChild<Run>();
                    if (run != null) {
                        var text = run.GetFirstChild<Text>();
                        if (text != null) {
                            return text.Text;
                        }
                    }
                }
                return "";
            }
            set {
                if (_sdtBlock != null) {
                    var run = _sdtBlock.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Run>();
                    if (run != null) {
                        var text = run.GetFirstChild<Text>();
                        if (text != null) {
                            text.Text = value;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Allows to modify the paragraph of the text box, along with other text formatting
        /// </summary>
        public WordParagraph WordParagraph {
            get {
                if (_sdtBlock != null) {
                    var run = _sdtBlock.GetFirstChild<Run>();
                    return new WordParagraph(_document, _sdtBlock, run);
                }
                return null;
            }
        }

        public HorizontalRelativePositionValues? HorizontalPositionRelativeFrom {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var horizontalPosition = anchor.HorizontalPosition;
                    if (horizontalPosition != null) {
                        return horizontalPosition.RelativeFrom;
                    }
                }

                return null;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    var horizontalPosition = anchor.HorizontalPosition;
                    if (horizontalPosition != null) {
                        horizontalPosition.RelativeFrom = value;
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the wrap text of the text box
        /// </summary>
        public WrapTextImage? WrapText {
            get => WordWrapTextImage.GetWrapTextImage(_anchor, _inline);
            set => WordWrapTextImage.SetWrapTextImage(_drawing, _anchor, _inline, value);
        }

        public DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues HorizontalAlignment {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var horizontalPosition = anchor.HorizontalPosition;
                    if (horizontalPosition != null && horizontalPosition.HorizontalAlignment != null) {
                        return GetHorizontalAlignmentFromText(horizontalPosition.HorizontalAlignment.Text);
                    }
                }
                return DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues.Center;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    var horizontalPosition = anchor.HorizontalPosition;
                    if (horizontalPosition == null) {
                        horizontalPosition = AddHorizontalPosition(anchor, true);
                    }
                    if (horizontalPosition.HorizontalAlignment == null) {
                        horizontalPosition.HorizontalAlignment = new HorizontalAlignment() {
                            Text = value.ToString().ToLower()
                        };
                    } else {
                        horizontalPosition.HorizontalAlignment.Text = value.ToString().ToLower();
                    }
                }
            }
        }

        public VerticalRelativePositionValues VerticalPositionRelativeFrom {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var verticalPosition = anchor.VerticalPosition;
                    if (verticalPosition != null) {
                        return verticalPosition.RelativeFrom;
                    }
                }
                return VerticalRelativePositionValues.Page;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    var verticalPosition = anchor.VerticalPosition;
                    if (verticalPosition != null) {
                        verticalPosition.RelativeFrom = value;
                    }
                }
            }
        }

        /// <summary>
        /// Allows to set vertically position of the text box in twips (twentieths of a point)
        /// </summary>
        public int? VerticalPositionOffset {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var verticalPosition = anchor.VerticalPosition;
                    if (verticalPosition != null) {
                        return int.Parse(verticalPosition.PositionOffset.Text);
                    }
                }

                return null;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    var verticalPosition = AddVerticalPosition(anchor, true);
                    if (verticalPosition != null) {
                        verticalPosition.PositionOffset.Text = value.ToString();
                    }
                }
            }
        }

        /// <summary>
        /// Allows to set vertically position of the text box in twips (twentieths of a point)
        /// Please remember that this property will remove alignment of the text box and instead use Absolute position
        /// </summary>
        public int? HorizonalPositionOffset {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var horizontalPosition = anchor.HorizontalPosition;
                    if (horizontalPosition != null && horizontalPosition.PositionOffset != null) {
                        return int.Parse(horizontalPosition.PositionOffset.Text);
                    }
                }
                return null;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    var horizontalPosition = AddHorizontalPosition(anchor, true);
                    if (horizontalPosition != null) {
                        horizontalPosition.PositionOffset.Text = value.ToString();
                    }
                }
            }
        }

        /// <summary>
        /// Allows to set horizontally position of the text box in centimeters
        /// Please 
        /// </summary>
        public double? HorizonalPositionOffsetCentimeters {
            get {
                if (HorizonalPositionOffset != null) {
                    return Helpers.ConvertEmusToCentimeters(HorizonalPositionOffset.Value);
                }

                return null;
            }
            set {
                if (value != null) {
                    HorizonalPositionOffset = Helpers.ConvertCentimetersToEmus(value.Value);
                }
            }
        }

        /// <summary>
        /// Allows to set vertically position of the text box in centimeters
        /// </summary>
        public double? VerticalPositionOffsetCentimeters {
            get {
                if (VerticalPositionOffset != null) {
                    return Helpers.ConvertEmusToCentimeters(VerticalPositionOffset.Value);
                }

                return null;
            }

            set {
                if (value != null) {
                    VerticalPositionOffset = Helpers.ConvertCentimetersToEmus(value.Value);
                }
            }
        }

        public int? RelativeWidthPercentage {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var relativeWidth = anchor.ChildElements.OfType<DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeWidth>().FirstOrDefault();
                    if (relativeWidth != null) {
                        if (relativeWidth.PercentageWidth != null) {
                            return int.Parse(relativeWidth.PercentageWidth.Text) / 1000;
                        }
                    }
                }
                return null;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    if (value != null) {
                        var setValue = value.Value * 1000;

                        var relativeWidth = anchor.ChildElements.OfType<DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeWidth>().FirstOrDefault();
                        if (relativeWidth == null) {
                            relativeWidth = new DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeWidth() {
                                PercentageWidth = new DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageWidth() {
                                    Text = setValue.ToString()
                                }
                            };
                            anchor.Append(relativeWidth);
                        } else {
                            if (relativeWidth.PercentageWidth == null) {
                                relativeWidth.PercentageWidth = new DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageWidth() {
                                    Text = setValue.ToString()
                                };
                            } else {
                                relativeWidth.PercentageWidth.Text = setValue.ToString();
                            }
                        }
                    } else {
                        // value is null
                    }
                }
            }
        }

        public int? RelativeHeightPercentage {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var relativeHeight = anchor.ChildElements.OfType<DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeHeight>().FirstOrDefault();
                    if (relativeHeight != null) {
                        if (relativeHeight.PercentageHeight != null) {
                            return int.Parse(relativeHeight.PercentageHeight.Text) / 1000;
                        }
                    }
                }
                return null;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    if (value != null) {
                        var setValue = value.Value * 1000;

                        var relativeHeight = anchor.ChildElements.OfType<DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeHeight>().FirstOrDefault();
                        if (relativeHeight == null) {
                            relativeHeight = new DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeHeight() {
                                PercentageHeight = new DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageHeight() {
                                    Text = setValue.ToString()
                                }
                            };
                            anchor.Append(relativeHeight);
                        } else {
                            if (relativeHeight.PercentageHeight == null) {
                                relativeHeight.PercentageHeight = new DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageHeight() {
                                    Text = setValue.ToString()
                                };
                            } else {
                                relativeHeight.PercentageHeight.Text = setValue.ToString();
                            }
                        }
                    } else {
                        // value is null
                    }
                }
            }
        }

        public TextBodyProperties TextBodyProperties {
            get {
                return _wordprocessingShape.ChildElements.OfType<TextBodyProperties>().FirstOrDefault();
            }
        }

        public bool AutoFitToTextSize {
            get {
                return TextBodyProperties.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.ShapeAutoFit>().Any();
            }
            set {
                TextBodyProperties.RemoveChild(TextBodyProperties.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.ShapeAutoFit>().FirstOrDefault());
                if (value) {
                    TextBodyProperties.Append(new DocumentFormat.OpenXml.Drawing.ShapeAutoFit());
                } else {
                    TextBodyProperties.Append(new DocumentFormat.OpenXml.Drawing.NoAutoFit());
                }
            }
        }

        public DocumentFormat.OpenXml.Office2010.Word.Drawing.SizeRelativeHorizontallyValues? SizeRelativeHorizontally {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var relativeWidth = anchor.ChildElements.OfType<DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeWidth>().FirstOrDefault();
                    if (relativeWidth != null) {
                        if (relativeWidth.ObjectId != null) {
                            return relativeWidth.ObjectId;
                        }
                    }
                }
                return null;
            }
            set {

            }
        }

        public Int64 Width {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var extent = anchor.ChildElements.OfType<Extent>().FirstOrDefault();
                    if (extent != null) {
                        return Int64.Parse(extent.Cx);
                    }
                }
                return 0;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    var extent = anchor.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
                    if (extent == null) {
                        extent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() {
                            Cx = value,
                            Cy = 0L
                        };
                        anchor.Append(extent);
                    } else {
                        extent.Cx = value;
                    }
                }
            }
        }

        public Int64 Height {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var extent = anchor.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
                    if (extent != null) {
                        return Int64.Parse(extent.Cy);
                    }
                }
                return 0;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    var extent = anchor.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
                    if (extent == null) {
                        extent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() {
                            Cx = 0L,
                            Cy = value
                        };
                        anchor.Append(extent);
                    } else {
                        extent.Cy = value;
                    }
                }
            }
        }

        public double WidthCentimeters {
            get {
                return Helpers.ConvertEmusToCentimeters(Width);
            }
            set {
                Width = Helpers.ConvertCentimetersToEmusInt64(value);
            }
        }

        public double HeightCentimeters {
            get {
                return Helpers.ConvertEmusToCentimeters(Height);
            }
            set {
                Height = Helpers.ConvertCentimetersToEmusInt64(value);
            }
        }

        private Drawing _drawing {
            get {
                var alternateContent = _run.ChildElements.OfType<AlternateContent>().FirstOrDefault();
                if (alternateContent != null) {
                    var alternateContentChoice = alternateContent.ChildElements.OfType<AlternateContentChoice>().FirstOrDefault();
                    if (alternateContentChoice != null) {
                        var drawing = alternateContentChoice.ChildElements.OfType<DocumentFormat.OpenXml.Wordprocessing.Drawing>().FirstOrDefault();
                        if (drawing != null) {
                            return drawing;
                        }
                    }
                }

                return null;
            }
        }

        private Inline _inline {
            get {
                var alternateContent = _run.ChildElements.OfType<AlternateContent>().FirstOrDefault();
                if (alternateContent != null) {
                    var alternateContentChoice = alternateContent.ChildElements.OfType<AlternateContentChoice>().FirstOrDefault();
                    if (alternateContentChoice != null) {
                        var drawing = alternateContentChoice.ChildElements.OfType<DocumentFormat.OpenXml.Wordprocessing.Drawing>().FirstOrDefault();
                        if (drawing != null) {
                            var inline = drawing.Inline;
                            if (inline != null) {
                                return inline;
                            }
                        }
                    }
                }
                return null;
            }
        }

        private Anchor _anchor {
            get {
                var alternateContent = _run.ChildElements.OfType<AlternateContent>().FirstOrDefault();
                if (alternateContent != null) {
                    var alternateContentChoice = alternateContent.ChildElements.OfType<AlternateContentChoice>().FirstOrDefault();
                    if (alternateContentChoice != null) {
                        var drawing = alternateContentChoice.ChildElements.OfType<DocumentFormat.OpenXml.Wordprocessing.Drawing>().FirstOrDefault();
                        if (drawing != null) {
                            var anchor = drawing.Anchor;
                            if (anchor != null) {
                                return anchor;
                            }
                        }
                    }
                }
                return null;
            }
        }

        internal static Anchor ConvertInlineToAnchor(Inline inline, WrapTextImage wrapTextImage) {
            Anchor anchor1 = new Anchor() { DistanceFromTop = (UInt32Value)91440U, DistanceFromBottom = (UInt32Value)91440U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "39C62DE8", AnchorId = "3E379294" };
            anchor1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            anchor1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            SimplePosition simplePosition1 = new SimplePosition() { X = 0L, Y = 0L };

            HorizontalPosition horizontalPosition1 = new HorizontalPosition() { RelativeFrom = HorizontalRelativePositionValues.Page };
            HorizontalAlignment horizontalAlignment1 = new HorizontalAlignment();
            horizontalAlignment1.Text = "center";
            horizontalPosition1.Append(horizontalAlignment1);

            VerticalPosition verticalPosition1 = new VerticalPosition() { RelativeFrom = VerticalRelativePositionValues.Page };
            PositionOffset positionOffset1 = new PositionOffset();
            positionOffset1.Text = "182880";
            verticalPosition1.Append(positionOffset1);

            // Copy INLINE to ANCHOR
            Extent extent1 = (Extent)inline.Extent.CloneNode(true);
            EffectExtent effectExtent1 = (EffectExtent)inline.EffectExtent.CloneNode(true);
            DocProperties docProperties1 = (DocProperties)inline.DocProperties.CloneNode(true);
            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = (NonVisualGraphicFrameDrawingProperties)inline.NonVisualGraphicFrameDrawingProperties.CloneNode(true);
            Graphic graphic1 = (Graphic)inline.Graphic.CloneNode(true);

            // continuation of anchor
            DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeWidth relativeWidth1 = new DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeWidth() { ObjectId = DocumentFormat.OpenXml.Office2010.Word.Drawing.SizeRelativeHorizontallyValues.Margin };
            DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageWidth percentageWidth1 = new DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageWidth();
            percentageWidth1.Text = "58500";

            relativeWidth1.Append(percentageWidth1);

            DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeHeight relativeHeight1 = new DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeHeight() { RelativeFrom = DocumentFormat.OpenXml.Office2010.Word.Drawing.SizeRelativeVerticallyValues.Margin };
            DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageHeight percentageHeight1 = new DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageHeight();
            percentageHeight1.Text = "20000";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);

            WordWrapTextImage.AppendWrapTextImage(anchor1, wrapTextImage);

            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);
            return anchor1;
        }

        internal static Inline ConvertAnchorToInline(Anchor anchor) {
            Inline inline1 = new Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "29D0141D", EditId = "5A52866D" };

            Extent extent1 = (Extent)anchor.Extent.CloneNode(true);
            EffectExtent effectExtent1 = (EffectExtent)anchor.EffectExtent.CloneNode(true);
            DocProperties docProperties1 = (DocProperties)anchor.OfType<DocProperties>().FirstOrDefault()?.CloneNode(true);
            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = (NonVisualGraphicFrameDrawingProperties)anchor.OfType<NonVisualGraphicFrameDrawingProperties>().FirstOrDefault()?.CloneNode(true);
            Graphic graphic1 = (Graphic)anchor.OfType<Graphic>().FirstOrDefault()?.CloneNode(true);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            return inline1;
        }

        private DocumentFormat.OpenXml.Drawing.GraphicData _graphicData {
            get {
                var graphic = _anchor.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Graphic>().FirstOrDefault();
                if (graphic != null) {
                    return graphic.GraphicData;
                }
                return null;
            }
        }

        private DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape _wordprocessingShape {
            get {
                var graphicData = _graphicData;
                if (graphicData != null) {
                    var wsp = graphicData.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape>();
                    if (wsp != null) {
                        return wsp;
                    }
                }
                return null;
            }
        }

        private Paragraph _sdtBlock {
            get {
                var wordprocessingShape = _wordprocessingShape;
                if (wordprocessingShape != null) {


                    var textBoxInfo = wordprocessingShape.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.TextBoxInfo2>();
                    if (textBoxInfo != null) {
                        var textBoxContent = textBoxInfo.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.TextBoxContent>();
                        if (textBoxContent != null) {
                            var sdtBlock = textBoxContent.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
                            if (sdtBlock != null) {
                                return sdtBlock;
                            }
                        }
                    }

                }
                return null;
            }
        }

        private SdtContentBlock _sdtContentBlock {
            get {
                var sdtBlock = _sdtBlock;
                if (sdtBlock != null) {
                    var sdtContentBlock = sdtBlock.GetFirstChild<SdtContentBlock>();
                    if (sdtContentBlock != null) {
                        return sdtContentBlock;
                    }
                }
                return null;
            }
        }

        private VerticalPosition AddVerticalPosition(Anchor anchor, bool expectedPositionOffset = false) {
            if (anchor != null) {
                var verticalPosition = anchor.VerticalPosition;
                if (verticalPosition == null) {
                    anchor.VerticalPosition = new VerticalPosition() {
                        RelativeFrom = VerticalRelativePositionValues.Page, VerticalAlignment = new VerticalAlignment() {
                            Text = "top"
                        }
                    };
                    verticalPosition = anchor.VerticalPosition;
                }

                if (expectedPositionOffset) {
                    var positionOffset = verticalPosition.PositionOffset;
                    if (positionOffset == null) {
                        verticalPosition.PositionOffset = new PositionOffset() {
                            Text = "0"
                        };
                    }
                }
                return verticalPosition;
            }
            return null;
        }

        /// <summary>
        /// Small helper to create horizontal position if it doesn't exist
        /// </summary>
        /// <param name="anchor"></param>
        /// <param name="expectedPositionOffset"></param>
        /// <returns></returns>
        private HorizontalPosition AddHorizontalPosition(Anchor anchor, bool expectedPositionOffset = false) {
            if (anchor != null) {
                var horizontalPosition = anchor.HorizontalPosition;
                if (horizontalPosition == null && expectedPositionOffset) {
                    // position offset and horizontal alignment don't play together
                    anchor.HorizontalPosition = new HorizontalPosition() {
                        RelativeFrom = HorizontalRelativePositionValues.Page,
                    };
                    horizontalPosition = anchor.HorizontalPosition;
                } else if (horizontalPosition == null) {
                    anchor.HorizontalPosition = new HorizontalPosition() {
                        RelativeFrom = HorizontalRelativePositionValues.Page,
                        HorizontalAlignment = new HorizontalAlignment() {
                            Text = "center"
                        }
                    };
                    horizontalPosition = anchor.HorizontalPosition;
                }
                if (expectedPositionOffset) {
                    var positionOffset = horizontalPosition.PositionOffset;
                    if (positionOffset == null) {
                        positionOffset = new PositionOffset() {
                            Text = "0"
                        };
                        horizontalPosition.Append(positionOffset);
                    }
                    // we need to remove horizontal alignment if we want to use position offset
                    if (horizontalPosition.HorizontalAlignment != null) {
                        horizontalPosition.HorizontalAlignment.Remove();
                    }
                }
                return horizontalPosition;
            }
            return null;
        }

        private void AddAlternateContent(WordDocument wordDocument, WordParagraph wordParagraph, string text, WrapTextImage wrapTextImage) {

            AlternateContent alternateContent1 = new AlternateContent();
            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

            DocumentFormat.OpenXml.Wordprocessing.Drawing drawing1;
            if (wrapTextImage == WrapTextImage.InLineWithText) {
                // inline
                drawing1 = new DocumentFormat.OpenXml.Wordprocessing.Drawing {
                    Inline = GenerateInline(text)
                };
            } else {
                // anchor
                drawing1 = new DocumentFormat.OpenXml.Wordprocessing.Drawing {
                    Anchor = GenerateAnchor(text, wrapTextImage)
                };
            }
            alternateContentChoice1.Append(drawing1);

            alternateContent1.Append(alternateContentChoice1);
            //alternateContent1.Append(alternateContentFallback1);
            wordParagraph._run.Append(alternateContent1);
        }

        private Inline GenerateInline(string text) {
            Inline inline1 = new Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "29D0141D", EditId = "5A52866D" };

            Extent extent1 = GenerateExtent();
            EffectExtent effectExtent1 = GenerateEffectExtent();
            DocProperties docProperties1 = GenerateDocProperties();
            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = GenerateNonVisualGraphicFrameDrawingProperties();
            Graphic graphic1 = GenerateGraphic(text);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            return inline1;
        }


        private Extent GenerateExtent() {
            Extent extent1 = new Extent() { Cx = 2360930L, Cy = 1404620L };
            return extent1;
        }

        private EffectExtent GenerateEffectExtent() {
            EffectExtent effectExtent1 = new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            return effectExtent1;
        }

        private DocProperties GenerateDocProperties() {
            DocProperties docProperties1 = new DocProperties() { Id = (UInt32Value)307U, Name = "Text Box 2" };
            return docProperties1;
        }

        private Graphic GenerateGraphic(string text) {
            DocumentFormat.OpenXml.Drawing.Graphic graphic1 = new DocumentFormat.OpenXml.Drawing.Graphic();

            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            DocumentFormat.OpenXml.Drawing.GraphicData graphicData1 = new DocumentFormat.OpenXml.Drawing.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            WordprocessingShape wordprocessingShape1 = new WordprocessingShape();
            wordprocessingShape1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new NonVisualDrawingShapeProperties() { TextBox = true };
            DocumentFormat.OpenXml.Drawing.ShapeLocks shapeLocks1 = new DocumentFormat.OpenXml.Drawing.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties1.Append(shapeLocks1);

            ShapeProperties shapeProperties1 = GenerateShapeProperties();
            TextBoxInfo2 textBoxInfo21 = new TextBoxInfo2();

            var txtBoxContent = GenerateTextBoxContent(text);
            textBoxInfo21.Append(txtBoxContent);

            TextBodyProperties textBodyProperties1 = new TextBodyProperties() { Rotation = 0, Vertical = DocumentFormat.OpenXml.Drawing.TextVerticalValues.Horizontal, Wrap = DocumentFormat.OpenXml.Drawing.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues.Top, AnchorCenter = false };
            DocumentFormat.OpenXml.Drawing.ShapeAutoFit shapeAutoFit1 = new DocumentFormat.OpenXml.Drawing.ShapeAutoFit();

            textBodyProperties1.Append(shapeAutoFit1);

            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);

            wordprocessingShape1.Append(textBoxInfo21);
            wordprocessingShape1.Append(textBodyProperties1);

            graphicData1.Append(wordprocessingShape1);

            graphic1.Append(graphicData1);

            return graphic1;
        }

        private NonVisualGraphicFrameDrawingProperties GenerateNonVisualGraphicFrameDrawingProperties() {
            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new NonVisualGraphicFrameDrawingProperties();

            DocumentFormat.OpenXml.Drawing.GraphicFrameLocks graphicFrameLocks1 = new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks();
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);
            return nonVisualGraphicFrameDrawingProperties1;
        }

        private Anchor GenerateAnchor(string text, WrapTextImage wrapTextImage) {
            Anchor anchor1 = new Anchor() { DistanceFromTop = (UInt32Value)91440U, DistanceFromBottom = (UInt32Value)91440U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "39C62DE8", AnchorId = "3E379294" };
            anchor1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            anchor1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            SimplePosition simplePosition1 = new SimplePosition() { X = 0L, Y = 0L };

            HorizontalPosition horizontalPosition1 = new HorizontalPosition() { RelativeFrom = HorizontalRelativePositionValues.Page };
            HorizontalAlignment horizontalAlignment1 = new HorizontalAlignment();
            horizontalAlignment1.Text = "center";

            horizontalPosition1.Append(horizontalAlignment1);

            VerticalPosition verticalPosition1 = new VerticalPosition() { RelativeFrom = VerticalRelativePositionValues.Page };
            PositionOffset positionOffset1 = new PositionOffset();
            positionOffset1.Text = "182880";

            verticalPosition1.Append(positionOffset1);

            // same content as per inline
            Extent extent1 = GenerateExtent();
            EffectExtent effectExtent1 = GenerateEffectExtent();
            DocProperties docProperties1 = GenerateDocProperties();
            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = GenerateNonVisualGraphicFrameDrawingProperties();
            Graphic graphic1 = GenerateGraphic(text);

            // continuation of anchor
            DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeWidth relativeWidth1 = new DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeWidth() { ObjectId = DocumentFormat.OpenXml.Office2010.Word.Drawing.SizeRelativeHorizontallyValues.Margin };
            DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageWidth percentageWidth1 = new DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageWidth();
            percentageWidth1.Text = "58500";

            relativeWidth1.Append(percentageWidth1);

            DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeHeight relativeHeight1 = new DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeHeight() { RelativeFrom = DocumentFormat.OpenXml.Office2010.Word.Drawing.SizeRelativeVerticallyValues.Margin };
            DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageHeight percentageHeight1 = new DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageHeight();
            percentageHeight1.Text = "20000";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);

            WordWrapTextImage.AppendWrapTextImage(anchor1, wrapTextImage);

            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);
            return anchor1;
        }

        public TextBoxContent GenerateTextBoxContent(string text) {
            TextBoxContent textBoxContent1 = new TextBoxContent();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00000000", RsidRunAdditionDefault = "006713BC", ParagraphId = "100FFE99", TextId = "27C5287F" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();

            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();
            RunProperties runProperties1 = new RunProperties();
            run1.Append(runProperties1);

            foreach (var part in text.Split(new string[] { Environment.NewLine, "\r\n", "\n", "\r" }, StringSplitOptions.None)) {
                Text textPart = new Text();
                textPart.Text = part;
                run1.Append(textPart);
                run1.Append(new Break());
            }

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            textBoxContent1.Append(paragraph1);
            return textBoxContent1;
        }

        /// <summary>
        /// Helps to translate text to HorizontalAlignment
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues GetHorizontalAlignmentFromText(string text) {
            switch (text.ToLower()) {
                case "left":
                    return DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues.Left;
                case "right":
                    return DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues.Right;
                case "center":
                    return DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues.Center;
                case "outside":
                    return DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues.Outside;
                default:
                    return DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues.Center;
            }
        }


        private ShapeProperties GenerateShapeProperties() {
            ShapeProperties shapeProperties1 = new ShapeProperties() { BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto };

            DocumentFormat.OpenXml.Drawing.Transform2D transform2D1 = new DocumentFormat.OpenXml.Drawing.Transform2D();
            DocumentFormat.OpenXml.Drawing.Offset offset1 = new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L };
            DocumentFormat.OpenXml.Drawing.Extents extents1 = new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 2360930L, Cy = 1404620L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            DocumentFormat.OpenXml.Drawing.PresetGeometry presetGeometry1 = new DocumentFormat.OpenXml.Drawing.PresetGeometry() { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle };
            DocumentFormat.OpenXml.Drawing.AdjustValueList adjustValueList1 = new DocumentFormat.OpenXml.Drawing.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill1 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex1 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill1.Append(rgbColorModelHex1);

            DocumentFormat.OpenXml.Drawing.Outline outline1 = new DocumentFormat.OpenXml.Drawing.Outline() { Width = 9525 };

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill2 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex2 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "000000" };

            solidFill2.Append(rgbColorModelHex2);
            DocumentFormat.OpenXml.Drawing.Miter miter1 = new DocumentFormat.OpenXml.Drawing.Miter() { Limit = 800000 };
            DocumentFormat.OpenXml.Drawing.HeadEnd headEnd1 = new DocumentFormat.OpenXml.Drawing.HeadEnd();
            DocumentFormat.OpenXml.Drawing.TailEnd tailEnd1 = new DocumentFormat.OpenXml.Drawing.TailEnd();

            outline1.Append(solidFill2);
            outline1.Append(miter1);
            outline1.Append(headEnd1);
            outline1.Append(tailEnd1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            //shapeProperties1.Append(solidFill1);
            //shapeProperties1.Append(outline1);
            return shapeProperties1;
        }

        private ShapeStyle GenerateShapeStyle() {
            ShapeStyle shapeStyle1 = new ShapeStyle();

            DocumentFormat.OpenXml.Drawing.LineReference lineReference1 = new DocumentFormat.OpenXml.Drawing.LineReference() { Index = (UInt32Value)0U };
            DocumentFormat.OpenXml.Drawing.SchemeColor schemeColor1 = new DocumentFormat.OpenXml.Drawing.SchemeColor() { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent1 };

            lineReference1.Append(schemeColor1);

            DocumentFormat.OpenXml.Drawing.FillReference fillReference1 = new DocumentFormat.OpenXml.Drawing.FillReference() { Index = (UInt32Value)0U };
            DocumentFormat.OpenXml.Drawing.SchemeColor schemeColor2 = new DocumentFormat.OpenXml.Drawing.SchemeColor() { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor2);

            DocumentFormat.OpenXml.Drawing.EffectReference effectReference1 = new DocumentFormat.OpenXml.Drawing.EffectReference() { Index = (UInt32Value)0U };
            DocumentFormat.OpenXml.Drawing.SchemeColor schemeColor3 = new DocumentFormat.OpenXml.Drawing.SchemeColor() { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor3);

            DocumentFormat.OpenXml.Drawing.FontReference fontReference1 = new DocumentFormat.OpenXml.Drawing.FontReference() { Index = DocumentFormat.OpenXml.Drawing.FontCollectionIndexValues.Minor };
            DocumentFormat.OpenXml.Drawing.SchemeColor schemeColor4 = new DocumentFormat.OpenXml.Drawing.SchemeColor() { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Dark1 };

            fontReference1.Append(schemeColor4);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);
            return shapeStyle1;
        }


        public void Remove() {
            if (this._wordParagraph != null) {
                this._wordParagraph.Remove();
            }
        }
    }
}
