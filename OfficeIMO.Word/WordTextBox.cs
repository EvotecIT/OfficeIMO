using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    public class WordTextBox {
        private WordDocument _document;
        private WordParagraph _wordParagraph;
        private Run _run => _wordParagraph._run;

        /// <summary>
        /// Add a new text box to the document
        /// </summary>
        /// <param name="wordDocument"></param>
        /// <param name="text"></param>
        public WordTextBox(WordDocument wordDocument, string text) {
            var paragraph = new WordParagraph(wordDocument, true, true);
            wordDocument.AddParagraph(paragraph);
            paragraph._run.Append(new RunProperties());
            AddAlternateContent(wordDocument, paragraph, text);

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

        /// <summary>
        /// Allows to set the text of the text box
        /// For more advanced text formatting use WordParagraph property
        /// </summary>
        public string Text {
            get {
                if (_sdtBlock != null) {
                    var paragraph = _sdtContentBlock.GetFirstChild<W.Paragraph>();
                    if (paragraph != null) {
                        var run = paragraph.GetFirstChild<W.Run>();
                        if (run != null) {
                            var text = run.GetFirstChild<W.Text>();
                            if (text != null) {
                                return text.Text;
                            }
                        }
                    }
                }
                return "";
            }
            set {
                if (_sdtBlock != null) {
                    var paragraph = _sdtContentBlock.GetFirstChild<W.Paragraph>();
                    if (paragraph != null) {
                        var run = paragraph.GetFirstChild<W.Run>();
                        if (run != null) {
                            var text = run.GetFirstChild<W.Text>();
                            if (text != null) {
                                text.Text = value;
                            }
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
                if (_sdtContentBlock != null) {
                    var paragraph = _sdtContentBlock.GetFirstChild<W.Paragraph>();
                    if (paragraph != null) {
                        var run = paragraph.GetFirstChild<W.Run>();
                        return new WordParagraph(_document, paragraph, run);
                    }
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

        public Wp.HorizontalAlignmentValues HorizontalAlignment {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var horizontalPosition = anchor.HorizontalPosition;
                    if (horizontalPosition != null && horizontalPosition.HorizontalAlignment != null) {
                        return GetHorizontalAlignmentFromText(horizontalPosition.HorizontalAlignment.Text);
                    }
                }
                return Wp.HorizontalAlignmentValues.Center;
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
        /// Please remember that this property will remove alignment of the text box and instead use Absolute position
        /// </summary>
        public double? HorizonalPositionOffsetCentimeters {
            get {
                if (HorizonalPositionOffset != null) {
                    return ConvertTwipsToCentimeters(HorizonalPositionOffset.Value);
                }

                return null;
            }
            set {
                if (value != null) {
                    HorizonalPositionOffset = ConvertCentimetersToTwips(value.Value);
                }
            }
        }

        /// <summary>
        /// Allows to set vertically position of the text box in centimeters
        /// </summary>
        public double? VerticalPositionOffsetCentimeters {
            get {
                if (VerticalPositionOffset != null) {
                    return ConvertTwipsToCentimeters(VerticalPositionOffset.Value);
                }

                return null;
            }

            set {
                if (value != null) {
                    VerticalPositionOffset = ConvertCentimetersToTwips(value.Value);
                }
            }
        }

        public int? RelativeWidthPercentage {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var relativeWidth = anchor.ChildElements.OfType<Wp14.RelativeWidth>().FirstOrDefault();
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

                        var relativeWidth = anchor.ChildElements.OfType<Wp14.RelativeWidth>().FirstOrDefault();
                        if (relativeWidth == null) {
                            relativeWidth = new Wp14.RelativeWidth() {
                                PercentageWidth = new Wp14.PercentageWidth() {
                                    Text = setValue.ToString()
                                }
                            };
                            anchor.Append(relativeWidth);
                        } else {
                            if (relativeWidth.PercentageWidth == null) {
                                relativeWidth.PercentageWidth = new Wp14.PercentageWidth() {
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
                    var relativeHeight = anchor.ChildElements.OfType<Wp14.RelativeHeight>().FirstOrDefault();
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

                        var relativeHeight = anchor.ChildElements.OfType<Wp14.RelativeHeight>().FirstOrDefault();
                        if (relativeHeight == null) {
                            relativeHeight = new Wp14.RelativeHeight() {
                                PercentageHeight = new Wp14.PercentageHeight() {
                                    Text = setValue.ToString()
                                }
                            };
                            anchor.Append(relativeHeight);
                        } else {
                            if (relativeHeight.PercentageHeight == null) {
                                relativeHeight.PercentageHeight = new Wp14.PercentageHeight() {
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

        public Wp14.SizeRelativeHorizontallyValues? SizeRelativeHorizontally {
            get {
                var anchor = _anchor;
                if (anchor != null) {
                    var relativeWidth = anchor.ChildElements.OfType<Wp14.RelativeWidth>().FirstOrDefault();
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
                    var extent = anchor.ChildElements.OfType<Wp.Extent>().FirstOrDefault();
                    if (extent != null) {
                        return Int64.Parse(extent.Cx);
                    }
                }
                return 0;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    var extent = anchor.ChildElements.OfType<Wp.Extent>().FirstOrDefault();
                    if (extent == null) {
                        extent = new Wp.Extent() {
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
                    var extent = anchor.ChildElements.OfType<Wp.Extent>().FirstOrDefault();
                    if (extent != null) {
                        return Int64.Parse(extent.Cy);
                    }
                }
                return 0;
            }
            set {
                var anchor = _anchor;
                if (anchor != null) {
                    var extent = anchor.ChildElements.OfType<Wp.Extent>().FirstOrDefault();
                    if (extent == null) {
                        extent = new Wp.Extent() {
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
                return ConvertTwipsToCentimeters(Width);
            }
            set {
                Width = ConvertCentimetersToTwipsInt64(value);
            }
        }

        public double HeightCentimeters {
            get {
                return ConvertTwipsToCentimeters(Height);
            }
            set {
                Height = ConvertCentimetersToTwipsInt64(value);
            }
        }

        private Anchor _anchor {
            get {
                var alternateContent = _run.ChildElements.OfType<AlternateContent>().FirstOrDefault();
                if (alternateContent != null) {
                    var alternateContentChoice = alternateContent.ChildElements.OfType<AlternateContentChoice>().FirstOrDefault();
                    if (alternateContentChoice != null) {
                        var drawing = alternateContentChoice.ChildElements.OfType<W.Drawing>().FirstOrDefault();
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

        private DocumentFormat.OpenXml.Drawing.GraphicData _graphicData {
            get {
                var graphic = _anchor.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Graphic>().FirstOrDefault();
                if (graphic != null) {
                    return graphic.GraphicData;
                }
                return null;
            }
        }

        private Wps.WordprocessingShape _wordprocessingShape {
            get {
                var graphicData = _graphicData;
                if (graphicData != null) {
                    var wsp = graphicData.GetFirstChild<Wps.WordprocessingShape>();
                    if (wsp != null) {
                        return wsp;
                    }
                }
                return null;
            }
        }

        private SdtBlock _sdtBlock {
            get {
                var wordprocessingShape = _wordprocessingShape;
                if (wordprocessingShape != null) {


                    var textBoxInfo = wordprocessingShape.GetFirstChild<Wps.TextBoxInfo2>();
                    if (textBoxInfo != null) {
                        var textBoxContent = textBoxInfo.GetFirstChild<W.TextBoxContent>();
                        if (textBoxContent != null) {
                            var sdtBlock = textBoxContent.GetFirstChild<W.SdtBlock>();
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

        /// <summary>
        /// Converts centimeters to twips (twentieths of a point)
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private int? ConvertCentimetersToTwips(double value) {
            int twips = (int)(value * 360000);
            return twips;
        }

        /// <summary>
        /// Converts centimeters to twips (twentieths of a point) (Int64)
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private Int64 ConvertCentimetersToTwipsInt64(double value) {
            Int64 twips = (Int64)(value * 360000);
            return twips;
        }

        /// <summary>
        /// Converts twips (twentieths of a point) to centimeters
        /// </summary>
        /// <param name="horizonalPositionOffset"></param>
        /// <returns></returns>
        private double? ConvertTwipsToCentimeters(int twipsValue) {
            double centimeters = (double)((double)twipsValue / (double)360000);
            return centimeters;
        }

        private double ConvertEmuToCentimeters(Int64 emuValue) {
            double centimeters = (double)((double)emuValue / (double)914400);
            return centimeters;
        }

        private Int64 ConvertCentimetersToEmu(double value) {
            Int64 emu = (Int64)(value * 914400);
            return emu;
        }

        /// <summary>
        /// Converts twips (twentieths of a point) to centimeters (Int64)
        /// </summary>
        /// <param name="twipsValue"></param>
        /// <returns></returns>
        private double ConvertTwipsToCentimeters(Int64 twipsValue) {
            double centimeters = (double)((double)twipsValue / (double)360000);
            return centimeters;
        }

        private void AddAlternateContent(WordDocument wordDocument, WordParagraph wordParagraph, string text) {

            AlternateContent alternateContent1 = new AlternateContent();
            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

            W.Drawing drawing1 = new W.Drawing {
                Anchor = GenerateAnchor(text)
            };

            alternateContentChoice1.Append(drawing1);

            //AlternateContentFallback alternateContentFallback1 = GenerateAlternateContentFallback(text);

            alternateContent1.Append(alternateContentChoice1);
            //alternateContent1.Append(alternateContentFallback1);
            wordParagraph._run.Append(alternateContent1);
        }

        /// <summary>
        /// This part is available when Microsoft Word creates TextBox. Not sure if it is needed.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        //public AlternateContentFallback GenerateAlternateContentFallback(string text) {
        //    AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

        //    W.Picture picture1 = new W.Picture();

        //    V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
        //    shapetype1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        //    shapetype1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
        //    shapetype1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
        //    shapetype1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "3E379294"));
        //    V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
        //    V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

        //    shapetype1.Append(stroke1);
        //    shapetype1.Append(path1);

        //    V.Shape shape1 = new V.Shape() { Id = "Text Box 2", Style = "position:absolute;margin-left:0;margin-top:228.5pt;width:273.6pt;height:110.55pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:585;mso-height-percent:200;mso-wrap-distance-left:9pt;mso-wrap-distance-top:7.2pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:7.2pt;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:absolute;mso-position-vertical-relative:page;mso-width-percent:585;mso-height-percent:200;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top", OptionalString = "_x0000_s1026", Filled = false, Stroked = false, Type = "#_x0000_t202", EncodedPackage = "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\n" };
        //    shape1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
        //    shape1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");

        //    V.TextBox textBox1 = new V.TextBox() { Style = "mso-fit-shape-to-text:t" };

        //    W.TextBoxContent textBoxContent1 = new W.TextBoxContent();

        //    W.SdtBlock sdtBlock1 = new W.SdtBlock();

        //    W.SdtProperties sdtProperties1 = new W.SdtProperties();

        //    W.RunProperties runProperties1 = new W.RunProperties();
        //    W.Italic italic1 = new W.Italic();
        //    W.ItalicComplexScript italicComplexScript1 = new W.ItalicComplexScript();
        //    W.Color color1 = new W.Color() { Val = "156082", ThemeColor = W.ThemeColorValues.Accent1 };
        //    W.FontSize fontSize1 = new W.FontSize() { Val = "24" };
        //    W.FontSizeComplexScript fontSizeComplexScript1 = new W.FontSizeComplexScript() { Val = "24" };

        //    runProperties1.Append(italic1);
        //    runProperties1.Append(italicComplexScript1);
        //    runProperties1.Append(color1);
        //    runProperties1.Append(fontSize1);
        //    runProperties1.Append(fontSizeComplexScript1);
        //    W.SdtId sdtId1 = new W.SdtId() { Val = 1469011327 };
        //    W.TemporarySdt temporarySdt1 = new W.TemporarySdt();
        //    W.ShowingPlaceholder showingPlaceholder1 = new W.ShowingPlaceholder();

        //    W15.Appearance appearance1 = new W15.Appearance() { Val = W15.SdtAppearance.Hidden };
        //    appearance1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");

        //    sdtProperties1.Append(runProperties1);
        //    sdtProperties1.Append(sdtId1);
        //    sdtProperties1.Append(temporarySdt1);
        //    sdtProperties1.Append(showingPlaceholder1);
        //    sdtProperties1.Append(appearance1);

        //    W.SdtContentBlock sdtContentBlock1 = new W.SdtContentBlock();

        //    W.Paragraph paragraph1 = new W.Paragraph() { RsidParagraphAddition = "00B16DB6", RsidRunAdditionDefault = "00B16DB6", ParagraphId = "506E57D5", TextId = "77777777" };
        //    paragraph1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

        //    W.ParagraphProperties paragraphProperties1 = new W.ParagraphProperties();

        //    W.ParagraphBorders paragraphBorders1 = new W.ParagraphBorders();
        //    W.TopBorder topBorder1 = new W.TopBorder() { Val = W.BorderValues.Single, Color = "156082", ThemeColor = W.ThemeColorValues.Accent1, Size = (UInt32Value)24U, Space = (UInt32Value)8U };
        //    W.BottomBorder bottomBorder1 = new W.BottomBorder() { Val = W.BorderValues.Single, Color = "156082", ThemeColor = W.ThemeColorValues.Accent1, Size = (UInt32Value)24U, Space = (UInt32Value)8U };

        //    paragraphBorders1.Append(topBorder1);
        //    paragraphBorders1.Append(bottomBorder1);
        //    W.SpacingBetweenLines spacingBetweenLines1 = new W.SpacingBetweenLines() { After = "0" };

        //    W.ParagraphMarkRunProperties paragraphMarkRunProperties1 = new W.ParagraphMarkRunProperties();
        //    W.Italic italic2 = new W.Italic();
        //    W.ItalicComplexScript italicComplexScript2 = new W.ItalicComplexScript();
        //    W.Color color2 = new W.Color() { Val = "156082", ThemeColor = W.ThemeColorValues.Accent1 };
        //    W.FontSize fontSize2 = new W.FontSize() { Val = "24" };

        //    paragraphMarkRunProperties1.Append(italic2);
        //    paragraphMarkRunProperties1.Append(italicComplexScript2);
        //    paragraphMarkRunProperties1.Append(color2);
        //    paragraphMarkRunProperties1.Append(fontSize2);

        //    paragraphProperties1.Append(paragraphBorders1);
        //    paragraphProperties1.Append(spacingBetweenLines1);
        //    paragraphProperties1.Append(paragraphMarkRunProperties1);

        //    W.Run run1 = new W.Run();

        //    W.RunProperties runProperties2 = new W.RunProperties();
        //    W.Italic italic3 = new W.Italic();
        //    W.ItalicComplexScript italicComplexScript3 = new W.ItalicComplexScript();
        //    W.Color color3 = new W.Color() { Val = "156082", ThemeColor = W.ThemeColorValues.Accent1 };
        //    W.FontSize fontSize3 = new W.FontSize() { Val = "24" };
        //    W.FontSizeComplexScript fontSizeComplexScript2 = new W.FontSizeComplexScript() { Val = "24" };

        //    runProperties2.Append(italic3);
        //    runProperties2.Append(italicComplexScript3);
        //    runProperties2.Append(color3);
        //    runProperties2.Append(fontSize3);
        //    runProperties2.Append(fontSizeComplexScript2);
        //    W.Text text1 = new W.Text();
        //    text1.Text = text;

        //    run1.Append(runProperties2);
        //    run1.Append(text1);

        //    paragraph1.Append(paragraphProperties1);
        //    paragraph1.Append(run1);

        //    sdtContentBlock1.Append(paragraph1);

        //    sdtBlock1.Append(sdtProperties1);
        //    sdtBlock1.Append(sdtContentBlock1);

        //    textBoxContent1.Append(sdtBlock1);

        //    textBox1.Append(textBoxContent1);

        //    Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { Type = Wvml.WrapValues.TopAndBottom, AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };
        //    textWrap1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");

        //    shape1.Append(textBox1);
        //    shape1.Append(textWrap1);

        //    picture1.Append(shapetype1);
        //    picture1.Append(shape1);

        //    alternateContentFallback1.Append(picture1);
        //    return alternateContentFallback1;
        //}

        private Anchor GenerateAnchor(string text) {
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
            Extent extent1 = new Extent() { Cx = 2360930L, Cy = 1404620L };
            EffectExtent effectExtent1 = new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            WrapTopBottom wrapTopBottom1 = new WrapTopBottom();
            DocProperties docProperties1 = new DocProperties() { Id = (UInt32Value)307U, Name = "Text Box 2" };

            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
            wordprocessingShape1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties1.Append(shapeLocks1);

            Wps.ShapeProperties shapeProperties1 = GenerateShapeProperties();

            //A.Transform2D transform2D1 = new A.Transform2D();
            //A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            //A.Extents extents1 = new A.Extents() { Cx = 3474720L, Cy = 1403985L };

            //transform2D1.Append(offset1);
            //transform2D1.Append(extents1);

            //A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            //A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            //presetGeometry1.Append(adjustValueList1);
            //A.NoFill noFill1 = new A.NoFill();

            //A.Outline outline1 = new A.Outline() { Width = 9525 };
            //A.NoFill noFill2 = new A.NoFill();
            //A.Miter miter1 = new A.Miter() { Limit = 800000 };
            //A.HeadEnd headEnd1 = new A.HeadEnd();
            //A.TailEnd tailEnd1 = new A.TailEnd();

            //outline1.Append(noFill2);
            //outline1.Append(miter1);
            //outline1.Append(headEnd1);
            //outline1.Append(tailEnd1);

            //shapeProperties1.Append(transform2D1);
            //shapeProperties1.Append(presetGeometry1);
            //shapeProperties1.Append(noFill1);
            //shapeProperties1.Append(outline1);

            Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

            W.TextBoxContent textBoxContent1 = new W.TextBoxContent();

            W.SdtBlock sdtBlock1 = new W.SdtBlock();

            W.SdtProperties sdtProperties1 = new W.SdtProperties();

            W.RunProperties runProperties1 = new W.RunProperties();
            W.Italic italic1 = new W.Italic();
            W.ItalicComplexScript italicComplexScript1 = new W.ItalicComplexScript();
            W.Color color1 = new W.Color() { Val = "156082", ThemeColor = W.ThemeColorValues.Accent1 };
            W.FontSize fontSize1 = new W.FontSize() { Val = "24" };
            W.FontSizeComplexScript fontSizeComplexScript1 = new W.FontSizeComplexScript() { Val = "24" };

            runProperties1.Append(italic1);
            runProperties1.Append(italicComplexScript1);
            runProperties1.Append(color1);
            runProperties1.Append(fontSize1);
            runProperties1.Append(fontSizeComplexScript1);
            W.SdtId sdtId1 = new W.SdtId() { Val = 1469011327 };
            W.TemporarySdt temporarySdt1 = new W.TemporarySdt();
            W.ShowingPlaceholder showingPlaceholder1 = new W.ShowingPlaceholder();

            W15.Appearance appearance1 = new W15.Appearance() { Val = W15.SdtAppearance.Hidden };
            appearance1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");

            sdtProperties1.Append(runProperties1);
            sdtProperties1.Append(sdtId1);
            sdtProperties1.Append(temporarySdt1);
            sdtProperties1.Append(showingPlaceholder1);
            sdtProperties1.Append(appearance1);

            W.SdtContentBlock sdtContentBlock1 = new W.SdtContentBlock();

            W.Paragraph paragraph1 = new W.Paragraph() { RsidParagraphAddition = "00B16DB6", RsidRunAdditionDefault = "00B16DB6", ParagraphId = "506E57D5", TextId = "77777777" };
            paragraph1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            W.ParagraphProperties paragraphProperties1 = new W.ParagraphProperties();

            //W.ParagraphBorders paragraphBorders1 = new W.ParagraphBorders();
            //W.TopBorder topBorder1 = new W.TopBorder() { Val = W.BorderValues.Single, Color = "156082", ThemeColor = W.ThemeColorValues.Accent1, Size = (UInt32Value)24U, Space = (UInt32Value)8U };
            //W.BottomBorder bottomBorder1 = new W.BottomBorder() { Val = W.BorderValues.Single, Color = "156082", ThemeColor = W.ThemeColorValues.Accent1, Size = (UInt32Value)24U, Space = (UInt32Value)8U };

            //paragraphBorders1.Append(topBorder1);
            //paragraphBorders1.Append(bottomBorder1);
            W.SpacingBetweenLines spacingBetweenLines1 = new W.SpacingBetweenLines() { After = "0" };

            W.ParagraphMarkRunProperties paragraphMarkRunProperties1 = new W.ParagraphMarkRunProperties();
            W.Italic italic2 = new W.Italic();
            W.ItalicComplexScript italicComplexScript2 = new W.ItalicComplexScript();
            W.Color color2 = new W.Color() { Val = "156082", ThemeColor = W.ThemeColorValues.Accent1 };
            W.FontSize fontSize2 = new W.FontSize() { Val = "24" };

            paragraphMarkRunProperties1.Append(italic2);
            paragraphMarkRunProperties1.Append(italicComplexScript2);
            paragraphMarkRunProperties1.Append(color2);
            paragraphMarkRunProperties1.Append(fontSize2);

            //paragraphProperties1.Append(paragraphBorders1);
            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            W.Run run1 = new W.Run();

            W.RunProperties runProperties2 = new W.RunProperties();
            W.Italic italic3 = new W.Italic();
            W.ItalicComplexScript italicComplexScript3 = new W.ItalicComplexScript();
            W.Color color3 = new W.Color() { Val = "156082", ThemeColor = W.ThemeColorValues.Accent1 };
            W.FontSize fontSize3 = new W.FontSize() { Val = "24" };
            W.FontSizeComplexScript fontSizeComplexScript2 = new W.FontSizeComplexScript() { Val = "24" };

            runProperties2.Append(italic3);
            runProperties2.Append(italicComplexScript3);
            runProperties2.Append(color3);
            runProperties2.Append(fontSize3);
            runProperties2.Append(fontSizeComplexScript2);
            W.Text text1 = new W.Text();
            text1.Text = text;

            run1.Append(runProperties2);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            sdtContentBlock1.Append(paragraph1);

            sdtBlock1.Append(sdtProperties1);
            sdtBlock1.Append(sdtContentBlock1);

            textBoxContent1.Append(sdtBlock1);

            textBoxInfo21.Append(textBoxContent1);

            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false };
            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            textBodyProperties1.Append(shapeAutoFit1);

            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);

            //ShapeStyle shapeStyle1 = GenerateShapeStyle();
            //wordprocessingShape1.Append(shapeStyle1);


            wordprocessingShape1.Append(textBoxInfo21);
            wordprocessingShape1.Append(textBodyProperties1);

            graphicData1.Append(wordprocessingShape1);

            graphic1.Append(graphicData1);

            Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin };
            Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
            percentageWidth1.Text = "58500";

            relativeWidth1.Append(percentageWidth1);

            Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Margin };
            Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
            percentageHeight1.Text = "20000";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapTopBottom1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);
            return anchor1;
        }

        /// <summary>
        /// Helps to translate text to HorizontalAlignment
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private Wp.HorizontalAlignmentValues GetHorizontalAlignmentFromText(string text) {
            switch (text.ToLower()) {
                case "left":
                    return Wp.HorizontalAlignmentValues.Left;
                case "right":
                    return Wp.HorizontalAlignmentValues.Right;
                case "center":
                    return Wp.HorizontalAlignmentValues.Center;
                case "outside":
                    return Wp.HorizontalAlignmentValues.Outside;
                default:
                    return Wp.HorizontalAlignmentValues.Center;
            }
        }


        private ShapeProperties GenerateShapeProperties() {
            ShapeProperties shapeProperties1 = new ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 2360930L, Cy = 1404620L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill1.Append(rgbColorModelHex1);

            A.Outline outline1 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill2.Append(rgbColorModelHex2);
            A.Miter miter1 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd1 = new A.HeadEnd();
            A.TailEnd tailEnd1 = new A.TailEnd();

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

            A.LineReference lineReference1 = new A.LineReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            lineReference1.Append(schemeColor1);

            A.FillReference fillReference1 = new A.FillReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor2);

            A.EffectReference effectReference1 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor3);

            A.FontReference fontReference1 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference1.Append(schemeColor4);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);
            return shapeStyle1;
        }



    }
}
