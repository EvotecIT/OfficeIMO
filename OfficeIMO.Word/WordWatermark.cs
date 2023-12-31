using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocumentFormat.OpenXml.Vml;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    public enum WordWatermarkStyle {
        Text,
        Image
    }

    public class WordWatermark {
        private WordDocument _document;
        private SdtBlock _sdtBlock;
        private WordHeader _wordHeader;
        private WordSection _section;
        private WordParagraph _wordParagraph;

        public string Text {
            get {
                var paragraph = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>().FirstOrDefault();
                if (paragraph != null) {
                    var run = paragraph.Descendants().OfType<Run>().FirstOrDefault();
                    if (run != null) {
                        var picture = run.Descendants().OfType<Picture>().FirstOrDefault();
                        if (picture != null) {
                            var shape = picture.Descendants().OfType<Shape>().FirstOrDefault();
                            if (shape != null) {
                                TextPath textPath = shape.GetFirstChild<V.TextPath>();
                                if (textPath != null) {
                                    return textPath.String;

                                }
                            }
                        }
                    }
                }

                return "";
            }
            set {
                var paragraph = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>().FirstOrDefault();
                if (paragraph != null) {
                    var run = paragraph.Descendants().OfType<Run>().FirstOrDefault();
                    if (run != null) {
                        var picture = run.Descendants().OfType<Picture>().FirstOrDefault();
                        if (picture != null) {
                            var shape = picture.Descendants().OfType<Shape>().FirstOrDefault();
                            if (shape != null) {
                                TextPath textPath = shape.GetFirstChild<V.TextPath>();
                                if (textPath != null) {
                                    textPath.String = value;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Get or set rotation of the watermark.
        /// </summary>
        public int? Rotation {
            get {
                var shape = _shape;
                if (shape != null) {

                    // Style = "position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s1025", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
                    //
                    var style = shape.Style.Value;
                    if (style != null) {
                        var rotation = style.Split(';').FirstOrDefault(c => c.StartsWith("rotation:"));
                        if (rotation != null) {
                            var rotationValue = rotation.Split(':').LastOrDefault();
                            if (rotationValue != null) {
                                return int.Parse(rotationValue);
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                if (shape != null) {
                    var style = shape.Style.Value;
                    if (style != null) {
                        var rotation = style.Split(';').FirstOrDefault(c => c.StartsWith("rotation:"));
                        if (rotation != null) {
                            var rotationValue = rotation.Split(':').LastOrDefault();
                            if (rotationValue != null) {
                                shape.Style.Value = style.Replace(rotation, "rotation:" + value);
                            }
                        }
                    }
                }
            }
        }

        public double? Width {
            get {
                var shape = _shape;
                if (shape != null) {
                    var style = shape.Style.Value;
                    if (style != null) {
                        var width = style.Split(';').FirstOrDefault(c => c.StartsWith("width:"));
                        if (width != null) {
                            var widthValue = width.Split(':').LastOrDefault();
                            if (widthValue != null) {
                                string stringValue = widthValue.Replace("pt", "");
                                return double.Parse(stringValue, CultureInfo.InvariantCulture);
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                if (shape != null) {
                    var style = shape.Style.Value;
                    if (style != null) {
                        var width = style.Split(';').FirstOrDefault(c => c.StartsWith("width:"));
                        if (width != null) {
                            var widthValue = width.Split(':').LastOrDefault();
                            if (widthValue != null) {
                                shape.Style.Value = style.Replace(width, "width:" + value + "pt");
                            }
                        }
                    }
                }
            }
        }

        public double? Height {
            get {
                var shape = _shape;
                if (shape != null) {
                    var style = shape.Style.Value;
                    if (style != null) {
                        var height = style.Split(';').FirstOrDefault(c => c.StartsWith("height:"));
                        if (height != null) {
                            var heightValue = height.Split(':').LastOrDefault();
                            if (heightValue != null) {
                                string stringValue = heightValue.Replace("pt", "");
                                return double.Parse(stringValue, CultureInfo.InvariantCulture);
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                if (shape != null) {
                    var style = shape.Style.Value;
                    if (style != null) {
                        var height = style.Split(';').FirstOrDefault(c => c.StartsWith("height:"));
                        if (height != null) {
                            var heightValue = height.Split(':').LastOrDefault();
                            if (heightValue != null) {
                                shape.Style.Value = style.Replace(height, "height:" + value + "pt");
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Get or Set if watermark is stroked.
        /// </summary>
        public bool Stroked {
            get {
                var shape = _shape;
                if (shape != null) {
                    return shape.Stroked.Value;
                }
                return false;
            }
            set {
                var shape = _shape;
                if (shape != null) {
                    shape.Stroked.Value = value;
                }
            }
        }

        /// <summary>
        /// Get or Set if watermark is allowed in cell.
        /// </summary>
        public bool? AllowInCell {
            get {
                var shape = _shape;
                if (shape != null) {
                    return shape.AllowInCell.Value;
                }
                return null;
            }
            set {
                var shape = _shape;
                if (shape != null) {
                    shape.AllowInCell.Value = value.Value;
                }
            }
        }

        /// <summary>
        /// Gets or sets color of the watermark.
        /// Some colors are not supported. If you set unsupported color, it will be ignored.
        /// </summary>
        public SixLabors.ImageSharp.Color? Color {
            get {
                if (ColorHex == "") {
                    return null;
                }

                return SixLabors.ImageSharp.Color.Parse(ColorHex);

            }
            set {
                if (value != null) {
                    this.ColorHex = ColorNameResolver.GetColorName(value.Value); //value.Value.ToHex();
                }
            }
        }

        /// <summary>
        /// Gets or sets color of the watermark using hex value.
        /// Setting colors using this property is not recommended because it doesn't actually accept hex values, but rather color names.
        /// </summary>
        public string ColorHex {
            get {
                var shape = _shape;
                if (shape != null && shape.FillColor != null) {
                    return shape.FillColor.Value;
                }
                return "";
            }
            set {
                var shape = _shape;
                if (shape != null) {
                    shape.FillColor.Value = value;
                }
            }
        }

        private Shape _shape {
            get {

                var shape = _picture.Descendants().OfType<Shape>().FirstOrDefault();
                if (shape != null) {
                    return shape;
                }
                return null;
            }
        }

        private Picture _picture {
            get {
                if (_sdtBlock != null) {
                    if (_sdtBlock.SdtContentBlock != null) {
                        var paragraph = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>().FirstOrDefault();
                        if (paragraph != null) {
                            var run = paragraph.Descendants().OfType<Run>().FirstOrDefault();
                            if (run != null) {
                                var picture = run.Descendants().OfType<Picture>().FirstOrDefault();
                                if (picture != null) {
                                    return picture;
                                }
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

        private static SdtBlock GetStyle(WordWatermarkStyle style) {
            switch (style) {
                case WordWatermarkStyle.Text: return TextWatermark;
                case WordWatermarkStyle.Image: return Confidential2;
            }

            throw new ArgumentOutOfRangeException(nameof(style));
        }

        public WordWatermark(WordDocument wordDocument, WordSection wordSection, WordWatermarkStyle style, string text) {
            this._document = wordDocument;
            this._section = wordSection;

            if (style == WordWatermarkStyle.Text) {
                this._sdtBlock = GetStyle(style);

                this.Text = text;

                //this._document._document.Body.Append(_sdtBlock);
                if (this._section.Paragraphs.Count == 0) {
                    this._document._document.Body.Append(_sdtBlock);
                } else {
                    var lastParagraph = this._section.Paragraphs.Last();
                    lastParagraph._paragraph.Parent.Append(_sdtBlock);
                }
            }
        }

        public WordWatermark(WordDocument wordDocument, WordSection wordSection, WordHeader wordHeader, WordWatermarkStyle style, string text) {
            this._document = wordDocument;
            this._section = wordSection;

            if (wordHeader == null) {
                // user didn't create headers first, so we do it for the user
                wordDocument.AddHeadersAndFooters();
                wordHeader = wordDocument.Header.Default;
            }
            this._wordHeader = wordHeader;

            if (style == WordWatermarkStyle.Text) {
                this._sdtBlock = GetStyle(style);


                this.Text = text;

                wordHeader._header.Append(_sdtBlock);
            }
        }

        public WordWatermark(WordDocument wordDocument, SdtBlock sdtBlock) {
            _document = wordDocument;
            _sdtBlock = sdtBlock;
        }

        private static SdtBlock TextWatermark {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -78212419 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Watermarks" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "003C040D", RsidRunAdditionDefault = "003C040D", ParagraphId = "7710D5F9", TextId = "47C6A96F" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "7FF27861" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@7,l@8,m@5,21600l@6,21600e" };

                V.Formulas formulas1 = new V.Formulas();
                V.Formula formula1 = new V.Formula() { Equation = "sum #0 0 10800" };
                V.Formula formula2 = new V.Formula() { Equation = "prod #0 2 1" };
                V.Formula formula3 = new V.Formula() { Equation = "sum 21600 0 @1" };
                V.Formula formula4 = new V.Formula() { Equation = "sum 0 0 @2" };
                V.Formula formula5 = new V.Formula() { Equation = "sum 21600 0 @3" };
                V.Formula formula6 = new V.Formula() { Equation = "if @0 @3 0" };
                V.Formula formula7 = new V.Formula() { Equation = "if @0 21600 @1" };
                V.Formula formula8 = new V.Formula() { Equation = "if @0 0 @2" };
                V.Formula formula9 = new V.Formula() { Equation = "if @0 @4 21600" };
                V.Formula formula10 = new V.Formula() { Equation = "mid @5 @6" };
                V.Formula formula11 = new V.Formula() { Equation = "mid @8 @5" };
                V.Formula formula12 = new V.Formula() { Equation = "mid @7 @8" };
                V.Formula formula13 = new V.Formula() { Equation = "mid @6 @7" };
                V.Formula formula14 = new V.Formula() { Equation = "sum @6 0 @5" };

                formulas1.Append(formula1);
                formulas1.Append(formula2);
                formulas1.Append(formula3);
                formulas1.Append(formula4);
                formulas1.Append(formula5);
                formulas1.Append(formula6);
                formulas1.Append(formula7);
                formulas1.Append(formula8);
                formulas1.Append(formula9);
                formulas1.Append(formula10);
                formulas1.Append(formula11);
                formulas1.Append(formula12);
                formulas1.Append(formula13);
                formulas1.Append(formula14);
                V.Path path1 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" };
                V.TextPath textPath1 = new V.TextPath() { On = true, FitShape = true };

                V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
                V.ShapeHandle shapeHandle1 = new V.ShapeHandle() { Position = "#0,bottomRight", XRange = "6629,14971" };

                shapeHandles1.Append(shapeHandle1);
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

                shapetype1.Append(formulas1);
                shapetype1.Append(path1);
                shapetype1.Append(textPath1);
                shapetype1.Append(shapeHandles1);
                shapetype1.Append(lock1);

                V.Shape shape1 = new V.Shape() { Id = "PowerPlusWaterMarkObject357476642", Style = "position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:90;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s1025", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
                V.Fill fill1 = new V.Fill() { Opacity = ".5" };
                V.TextPath textPath2 = new V.TextPath() { Style = "font-family:\"Calibri\";font-size:1pt;", String = "Draft" };
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };


                //V.Shape shape1 = new V.Shape() {Id = "PowerPlusWaterMarkObject357533252", Style = "position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s1028", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136"};
                //V.Fill fill1 = new V.Fill() {Opacity = ".5"};
                //V.TextPath textPath2 = new V.TextPath() {Style = "font-family:\"Calibri\";font-size:1pt", String = "Test"};
                //Wvml.TextWrap textWrap1 = new Wvml.TextWrap() {AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin};

                shape1.Append(fill1);
                shape1.Append(textPath2);
                shape1.Append(textWrap1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;
            }
        }

        private static SdtBlock Confidential2 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 1122028455 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Watermarks" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "003C040D", RsidRunAdditionDefault = "00F42210", ParagraphId = "7710D5F9", TextId = "7F2AA104" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "006695D0" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@7,l@8,m@5,21600l@6,21600e" };

                V.Formulas formulas1 = new V.Formulas();
                V.Formula formula1 = new V.Formula() { Equation = "sum #0 0 10800" };
                V.Formula formula2 = new V.Formula() { Equation = "prod #0 2 1" };
                V.Formula formula3 = new V.Formula() { Equation = "sum 21600 0 @1" };
                V.Formula formula4 = new V.Formula() { Equation = "sum 0 0 @2" };
                V.Formula formula5 = new V.Formula() { Equation = "sum 21600 0 @3" };
                V.Formula formula6 = new V.Formula() { Equation = "if @0 @3 0" };
                V.Formula formula7 = new V.Formula() { Equation = "if @0 21600 @1" };
                V.Formula formula8 = new V.Formula() { Equation = "if @0 0 @2" };
                V.Formula formula9 = new V.Formula() { Equation = "if @0 @4 21600" };
                V.Formula formula10 = new V.Formula() { Equation = "mid @5 @6" };
                V.Formula formula11 = new V.Formula() { Equation = "mid @8 @5" };
                V.Formula formula12 = new V.Formula() { Equation = "mid @7 @8" };
                V.Formula formula13 = new V.Formula() { Equation = "mid @6 @7" };
                V.Formula formula14 = new V.Formula() { Equation = "sum @6 0 @5" };

                formulas1.Append(formula1);
                formulas1.Append(formula2);
                formulas1.Append(formula3);
                formulas1.Append(formula4);
                formulas1.Append(formula5);
                formulas1.Append(formula6);
                formulas1.Append(formula7);
                formulas1.Append(formula8);
                formulas1.Append(formula9);
                formulas1.Append(formula10);
                formulas1.Append(formula11);
                formulas1.Append(formula12);
                formulas1.Append(formula13);
                formulas1.Append(formula14);
                V.Path path1 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" };
                V.TextPath textPath1 = new V.TextPath() { On = true, FitShape = true };

                V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
                V.ShapeHandle shapeHandle1 = new V.ShapeHandle() { Position = "#0,bottomRight", XRange = "6629,14971" };

                shapeHandles1.Append(shapeHandle1);
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

                shapetype1.Append(formulas1);
                shapetype1.Append(path1);
                shapetype1.Append(textPath1);
                shapetype1.Append(shapeHandles1);
                shapetype1.Append(lock1);

                V.Shape shape1 = new V.Shape() { Id = "PowerPlusWaterMarkObject357533252", Style = "position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s1028", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
                V.Fill fill1 = new V.Fill() { Opacity = ".5" };
                V.TextPath textPath2 = new V.TextPath() { Style = "font-family:\"Calibri\";font-size:1pt", String = "Test" };
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

                shape1.Append(fill1);
                shape1.Append(textPath2);
                shape1.Append(textWrap1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;
            }
        }

        private Paragraph ImageWatermark {
            get {
                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00083637", RsidRunAdditionDefault = "00083637", ParagraphId = "1543700F", TextId = "24264FF5" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "6CD641E4" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t75", CoordinateSize = "21600,21600", Filled = false, Stroked = false, OptionalNumber = 75, PreferRelative = true, EdgePath = "m@4@5l@4@11@9@11@9@5xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };

                V.Formulas formulas1 = new V.Formulas();
                V.Formula formula1 = new V.Formula() { Equation = "if lineDrawn pixelLineWidth 0" };
                V.Formula formula2 = new V.Formula() { Equation = "sum @0 1 0" };
                V.Formula formula3 = new V.Formula() { Equation = "sum 0 0 @1" };
                V.Formula formula4 = new V.Formula() { Equation = "prod @2 1 2" };
                V.Formula formula5 = new V.Formula() { Equation = "prod @3 21600 pixelWidth" };
                V.Formula formula6 = new V.Formula() { Equation = "prod @3 21600 pixelHeight" };
                V.Formula formula7 = new V.Formula() { Equation = "sum @0 0 1" };
                V.Formula formula8 = new V.Formula() { Equation = "prod @6 1 2" };
                V.Formula formula9 = new V.Formula() { Equation = "prod @7 21600 pixelWidth" };
                V.Formula formula10 = new V.Formula() { Equation = "sum @8 21600 0" };
                V.Formula formula11 = new V.Formula() { Equation = "prod @7 21600 pixelHeight" };
                V.Formula formula12 = new V.Formula() { Equation = "sum @10 21600 0" };

                formulas1.Append(formula1);
                formulas1.Append(formula2);
                formulas1.Append(formula3);
                formulas1.Append(formula4);
                formulas1.Append(formula5);
                formulas1.Append(formula6);
                formulas1.Append(formula7);
                formulas1.Append(formula8);
                formulas1.Append(formula9);
                formulas1.Append(formula10);
                formulas1.Append(formula11);
                formulas1.Append(formula12);
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle, AllowExtrusion = false };
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = true };

                shapetype1.Append(stroke1);
                shapetype1.Append(formulas1);
                shapetype1.Append(path1);
                shapetype1.Append(lock1);

                V.Shape shape1 = new V.Shape() { Id = "WordPictureWatermark269216046", Style = "position:absolute;margin-left:0;margin-top:0;width:467.3pt;height:148.7pt;z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s1025", AllowInCell = false, Type = "#_x0000_t75" };
                V.ImageData imageData1 = new V.ImageData() { Gain = "19661f", BlackLevel = "22938f", Title = "Logo-evotec-012", RelationshipId = "rId1" };

                shape1.Append(imageData1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);
                return paragraph1;
            }
        }

        public void Remove() {
            _sdtBlock.Remove();
        }
    }
}
